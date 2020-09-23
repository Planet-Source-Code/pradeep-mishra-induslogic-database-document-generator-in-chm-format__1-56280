VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TriggerObject As Trigger
Private m_TableObject As Table
Private strHTML As String
Private strTMPHTML As String

Public Property Get TriggerObject() As Trigger
    Set TriggerObject = m_TriggerObject
End Property

Public Property Set TriggerObject(ByVal vNewValue As Trigger)
    Set m_TriggerObject = vNewValue
End Property

Public Property Get TableObject() As Table
    Set TableObject = m_TableObject
End Property

Public Property Set TableObject(ByVal vNewValue As Table)
    Set m_TableObject = vNewValue
End Property

Public Function Header() As String
    strTMPHTML = m_TriggerObject.owner & "." & m_TriggerObject.Name
    strHTML = StartHead & StartTitle & "Trigger Details: " & strTMPHTML & EndTitle
    strHTML = strHTML & Replace(StartDIV, "<%HEADING%>", "Trigger Details :" & strTMPHTML) & "<BR>"
    Header = strHTML
End Function

Public Function Properties() As String
    strHTML = PropTriggerHeadHTML
    
    strHTML = Replace(strHTML, "<%TABLELINK%>", "Table." & m_TableObject.owner & "." & m_TableObject.Name & ".htm")
    strHTML = Replace(strHTML, "<%TABLENAME%>", m_TableObject.Name)
    strHTML = Replace(strHTML, "<%CREATIONDATE%>", m_TriggerObject.CreateDate)
    strHTML = Replace(strHTML, "<%PRECISION%>", IIf(m_TriggerObject.Properties(1).Value = True, "After", "Before"))
    strHTML = Replace(strHTML, "<%INSERTTRIGGER%>", IIf(m_TriggerObject.Type = (SQLDMOTrig_Insert) Or m_TriggerObject.Type = (SQLDMOTrig_Insert + SQLDMOTrig_Delete) Or m_TriggerObject.Type = (SQLDMOTrig_Insert + SQLDMOTrig_Update) Or m_TriggerObject.Type = (SQLDMOTrig_All), "True", "False"))
    strHTML = Replace(strHTML, "<%UPDATETRIGGER%>", IIf(m_TriggerObject.Type = (SQLDMOTrig_Update) Or m_TriggerObject.Type = (SQLDMOTrig_Update + SQLDMOTrig_Delete) Or m_TriggerObject.Type = (SQLDMOTrig_Update + SQLDMOTrig_Insert) Or m_TriggerObject.Type = (SQLDMOTrig_All), "True", "False"))
    strHTML = Replace(strHTML, "<%DELETETRIGGER%>", IIf(m_TriggerObject.Type = (SQLDMOTrig_Delete) Or m_TriggerObject.Type = (SQLDMOTrig_Delete + SQLDMOTrig_Insert) Or m_TriggerObject.Type = (SQLDMOTrig_Delete + SQLDMOTrig_Update) Or m_TriggerObject.Type = (SQLDMOTrig_All), "True", "False"))
    Properties = strHTML & "<br>"
End Function

Public Function Documentation() As String
    Dim strDesc As String
    
    strDesc = GetTriggerDescription(m_TriggerObject)
    
    If strDesc <> "" Then
        strHTML = PropTriggerDescHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        Documentation = strHTML & "<br>"
    Else
        Documentation = ""
    End If

End Function

Public Function SourceText() As String
    Dim strDesc As String
    
    strDesc = Replace(m_TriggerObject.Text, vbCrLf, "<BR>")
    If strDesc <> "" Then
        strHTML = PropTriggerSourceHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        SourceText = strHTML & "<br>"
    Else
        SourceText = ""
    End If

End Function
