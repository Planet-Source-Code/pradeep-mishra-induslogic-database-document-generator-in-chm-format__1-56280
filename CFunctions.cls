VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_UserDefinedFunctionObject As UserDefinedFunction
Private strHTML As String
Private strTMPHTML As String

Public Property Get UserDefinedFunctionObject() As UserDefinedFunction
    Set UserDefinedFunctionObject = m_UserDefinedFunctionObject
End Property

Public Property Set UserDefinedFunctionObject(ByVal vNewValue As UserDefinedFunction)
    Set m_UserDefinedFunctionObject = vNewValue
End Property

Public Function Header() As String
    strTMPHTML = m_UserDefinedFunctionObject.owner & "." & m_UserDefinedFunctionObject.Name
    strHTML = StartHead & StartTitle & "Stored Procedure Details: " & strTMPHTML & EndTitle
    strHTML = strHTML & Replace(StartDIV, "<%HEADING%>", "UserDefinedFunction Details :" & strTMPHTML) & "<BR>"
    Header = strHTML
End Function

Public Function Properties() As String
    strHTML = PropUserDefinedFunctionHeadHTML
    
    strHTML = Replace(strHTML, "<%CREATIONDATE%>", m_UserDefinedFunctionObject.CreateDate)
    strHTML = Replace(strHTML, "<%ENCRYPTED%>", IIf(m_UserDefinedFunctionObject.Encrypted = True, "True", "False"))
    strHTML = Replace(strHTML, "<%ISDETERMINISTIC%>", IIf(m_UserDefinedFunctionObject.IsDeterministic = True, "True", "False"))
    strHTML = Replace(strHTML, "<%ISSCHEMABOUND%>", IIf(m_UserDefinedFunctionObject.IsSchemaBound = True, "True", "False"))
    strHTML = Replace(strHTML, "<%TYPE%>", GetUDFType(m_UserDefinedFunctionObject.Type))
    Properties = strHTML & "<br>"
End Function

Public Function Documentation() As String
    Dim strDesc As String
    
    strDesc = GetUserDefinedFunctionDescription(m_UserDefinedFunctionObject)
    
    If strDesc <> "" Then
        strHTML = PropUserDefinedFunctionDescHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        Documentation = strHTML & "<br>"
    Else
        Documentation = ""
    End If

End Function

Public Function Parameters() As String
    Dim objParam As QueryResults
    Dim strColumnHTML As String
    Dim i As Long
    
    strHTML = ""
    Set objParam = m_UserDefinedFunctionObject.EnumParameters()
    
    strHTML = PropUserDefinedFunctionParameterHTML
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                For i = 1 To .rows
                    strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                    strHTML = strHTML & "<TD>" & .GetColumnString(i, 1) & "</TD>"
                    strHTML = strHTML & "<TD>" & .GetColumnString(i, 2) & "</TD>"
                    strHTML = strHTML & "<TD>" & .GetColumnString(i, 3) & "</TD>"
                    strHTML = strHTML & "<TD>" & IIf(.GetColumnString(i, 5) = 1, "Output", "Input") & "</TD>"
                    strHTML = strHTML & "</TR>"
                Next
            End If
        End If
    End With
    
    Set objParam = m_UserDefinedFunctionObject.EnumOutputs()
    With objParam
        strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
        strHTML = strHTML & "<TD>@RETURN_VALUE</TD>"
        strHTML = strHTML & "<TD>" & .GetColumnString(1, 2) & "</TD>"
        strHTML = strHTML & "<TD>" & .GetColumnString(1, 3) & "</TD>"
        strHTML = strHTML & "<TD>Output</TD>"
        strHTML = strHTML & "</TR>"
    End With
    strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
    Parameters = strHTML
End Function

Public Function Dependant() As String
    Dim objParam As QueryResults
    Dim strColumnHTML As String
    
    Dim strName As String
    Dim strType As String
    Dim strOwner As String
    Dim i As Long
    
    strHTML = ""
    Set objParam = m_UserDefinedFunctionObject.EnumDependencies()
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropUserDefinedFunctionDependantHTML
                For i = 1 To .rows
                    strName = .GetColumnString(i, 2)
                    strType = GetObjDesc(.GetColumnString(i, 1))
                    strOwner = .GetColumnString(i, 3)
                    
                    strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                    strHTML = strHTML & "<TD><A HREF='" & ReturnFileName(strType, strOwner, strName) & "'>" & strName & "</A></TD>"
                    strHTML = strHTML & "<TD>" & strType & "</TD>"
                    strHTML = strHTML & "<TD>" & .GetColumnString(i, 4) & "</TD>"
                    strHTML = strHTML & "</TR>"
                Next
                strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
            End If
        End If
    End With
    Dependant = strHTML
End Function

Public Function SourceText() As String
    Dim strDesc As String
    
    strDesc = Replace(m_UserDefinedFunctionObject.Text, vbCrLf, "<BR>")
    If strDesc <> "" Then
        strHTML = PropUserDefinedFunctionSourceHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        SourceText = strHTML & "<br>"
    Else
        SourceText = ""
    End If

End Function


