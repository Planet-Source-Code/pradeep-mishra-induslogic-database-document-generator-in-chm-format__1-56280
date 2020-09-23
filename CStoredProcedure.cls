VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStoredProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_StoredProcedureObject As StoredProcedure
Private strHTML As String
Private strTMPHTML As String

Public Property Get StoredProcedureObject() As StoredProcedure
    Set StoredProcedureObject = m_StoredProcedureObject
End Property

Public Property Set StoredProcedureObject(ByVal vNewValue As StoredProcedure)
    Set m_StoredProcedureObject = vNewValue
End Property

Public Function Header() As String
    strTMPHTML = m_StoredProcedureObject.owner & "." & m_StoredProcedureObject.Name
    strHTML = StartHead & StartTitle & "Stored Procedure Details: " & strTMPHTML & EndTitle
    strHTML = strHTML & Replace(StartDIV, "<%HEADING%>", "StoredProcedure Details :" & strTMPHTML) & "<BR>"
    Header = strHTML
End Function

Public Function Properties() As String
    strHTML = PropStoredProcedureHeadHTML
    
    strHTML = Replace(strHTML, "<%CREATIONDATE%>", m_StoredProcedureObject.CreateDate)
    strHTML = Replace(strHTML, "<%ENCRYPTED%>", IIf(m_StoredProcedureObject.Properties(3).Value = True, "True", "False"))
    strHTML = Replace(strHTML, "<%STARTUPSTOREDPROCEDURE%>", IIf(m_StoredProcedureObject.Startup = True, "True", "False"))
    Properties = strHTML & "<br>"
End Function

Public Function Documentation() As String
    Dim strDesc As String
    
    strDesc = GetStoredProcedureDescription(m_StoredProcedureObject)
    
    If strDesc <> "" Then
        strHTML = PropStoredProcedureDescHTML
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
    Set objParam = m_StoredProcedureObject.EnumParameters()
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropStoredProcedureParameterHTML
                For i = 1 To .rows
                     strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                     strHTML = strHTML & "<TD>" & .GetColumnString(i, 1) & "</TD>"
                     strHTML = strHTML & "<TD>" & .GetColumnString(i, 2) & "</TD>"
                     strHTML = strHTML & "<TD>" & .GetColumnString(i, 3) & "</TD>"
                     strHTML = strHTML & "<TD>" & IIf(.GetColumnString(i, 5) = 1, "True", "False") & "</TD>"
                     strHTML = strHTML & "</TR>"
                Next
                strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
            End If
        End If
    End With
    
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
    Set objParam = m_StoredProcedureObject.EnumDependencies()
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropStoredProcedureDependantHTML
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
    
    strDesc = Replace(m_StoredProcedureObject.Text, vbCrLf, "<BR>")
    If strDesc <> "" Then
        strHTML = PropStoredProcedureSourceHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        SourceText = strHTML & "<br>"
    Else
        SourceText = ""
    End If

End Function

