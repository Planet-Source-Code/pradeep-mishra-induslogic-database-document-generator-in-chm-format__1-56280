VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DatabaseDetail()
    Dim oFileWriter As CFileWriter
    Dim strTabTOC As String
    Dim strHTML As String
    
    Set oFileWriter = New CFileWriter
    With oFileWriter
        strHTML = Replace(PropDatabaseDetails, "<%SERVERNAME%>", thisServer)
        strHTML = Replace(strHTML, "<%DATABASENAME%>", thisDB)
        strHTML = Replace(strHTML, "<%DATABASESYSTEM%>", mServer.VersionString)
        strHTML = Replace(strHTML, "<%DATABASEVERSION%>", mServer.VersionMajor)
        strHTML = Replace(strHTML, "<%RUNDATE%>", mServer.ServerTime)
        strHTML = Replace(strHTML, "<%CREATIONDATE%>", oDatabase.CreateDate)
        .FileData = StartHead & strHTML & EndHead
        .filename = "db_details.htm"
        .Path = frmSelectDirectory.txtDirectory
        .WriteToFile
    End With

End Function

Public Function TableDOC()
    Dim oTable As Table
    Dim oFileWriter As CFileWriter
    Dim strTabTOC As String
    Dim strAllTables As String
    Set oFileWriter = New CFileWriter
    Dim i As Long
    
    frmStartProcessing.pbCurrentObject.Min = 1
    frmStartProcessing.pbCurrentObject.Max = oDatabase.Tables.Count + 1
    frmStartProcessing.pbCurrentObject.Value = 1
    
    strAllTables = StartHead & PropAllTableHTML
    
    For Each oTable In oDatabase.Tables
        If oTable.SystemObject = False Then
            i = i + 1
            With oFileWriter
                .filename = "Table." & oTable.owner & "." & oTable.Name & ".htm"
                .Path = frmSelectDirectory.txtDirectory
                .FileData = CreateTableHTMLDoc(oTable)
                .WriteToFile
            
                strTabTOC = strTabTOC & TOCTableD
                strTabTOC = Replace(strTabTOC, "<%TABLENAME%>", oTable.Name)
                strTabTOC = Replace(strTabTOC, "<%TABLELINK%>", .filename)
            End With
            
            strAllTables = strAllTables & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
            strAllTables = strAllTables & "<TD><A Href='" & oFileWriter.filename & "'>" & oTable.Name & "</A></TD>"
            strAllTables = strAllTables & "<TD Align='Right'>" & oTable.rows & "</TD>"
            strAllTables = strAllTables & "<TD Align='Right'>" & oTable.DataSpaceUsed & "</TD>"
            strAllTables = strAllTables & "<TD Align='Right'>" & oTable.IndexSpaceUsed & "</TD>"
            strAllTables = strAllTables & "<TD Align='Right'>" & oTable.CreateDate & "</TD></TR>"
        End If
        
        frmStartProcessing.pbCurrentObject.Value = frmStartProcessing.pbCurrentObject.Value + 1
        DoEvents
    Next
    
    strAllTables = strAllTables & "</TBODY>" & EndHead
    With oFileWriter
        .Path = frmSelectDirectory.txtDirectory
        .filename = "Table." & MakeCompatibleFileName(oDatabase.owner) & ".AllTables" & ".htm"
        .FileData = strAllTables
        .WriteToFile
    End With
    
    strIndex = strIndex & strTabTOC
    TOC = TOC & strTabTOC
End Function



Private Function CreateTableHTMLDoc(ByRef dbObject As Table) As String
    Dim strHTML As String
    Dim oTable As New CTable
            
    With oTable
        Set .TableObject = dbObject
        strHTML = .Header
        strHTML = strHTML & .Properties
        strHTML = strHTML & .Documentation
        strHTML = strHTML & .Column
        strHTML = strHTML & .Indexes
        strHTML = strHTML & .Dependant
        strHTML = strHTML & .ReferencedTables
        strHTML = strHTML & .ReferencingTables
    End With
    
    strHTML = strHTML & EndHead
    CreateTableHTMLDoc = strHTML
End Function

Public Function TriggerDOC()
    Dim oTable As Table
    Dim oTrigger As Trigger
    Dim oFileWriter As CFileWriter
    Dim strTriggerTOC As String
    
    Set oFileWriter = New CFileWriter
    
    frmStartProcessing.pbCurrentObject.Min = 1
    frmStartProcessing.pbCurrentObject.Max = oDatabase.Tables.Count + 1
    frmStartProcessing.pbCurrentObject.Value = 1
    
    For Each oTable In oDatabase.Tables
        If oTable.SystemObject = False Then
            For Each oTrigger In oTable.Triggers
                With oFileWriter
                    .filename = "Trigger." & oTable.owner & "." & oTable.Name & "." & oTrigger.Name & ".htm"
                    .Path = frmSelectDirectory.txtDirectory
                    .FileData = CreateTriggerHTMLDoc(oTable, oTrigger)
                    .WriteToFile
                    
                    strTriggerTOC = strTriggerTOC & TOCTriggerD
                    strTriggerTOC = Replace(strTriggerTOC, "<%TRIGGERNAME%>", oTrigger.Name)
                    strTriggerTOC = Replace(strTriggerTOC, "<%TRIGGERLINK%>", .filename)
                End With
            Next
        End If
        
        frmStartProcessing.pbCurrentObject.Value = frmStartProcessing.pbCurrentObject.Value + 1
        DoEvents
    Next
    strIndex = strIndex & strTriggerTOC
    TOC = TOC & strTriggerTOC
End Function

Private Function CreateTriggerHTMLDoc(ByRef oTable As Table, ByRef oTrigger As Trigger) As String
    Dim ObjTrigger As CTrigger
    Dim strHTML As String
            
    Set ObjTrigger = New CTrigger
    With ObjTrigger
        Set .TriggerObject = oTrigger
        Set .TableObject = oTable
        
        strHTML = .Header
        strHTML = strHTML & .Properties
        strHTML = strHTML & .Documentation
        strHTML = strHTML & .SourceText
        strHTML = strHTML & EndHead
    End With
    CreateTriggerHTMLDoc = strHTML
End Function

Public Function ViewsDOC()
    Dim oView As View
    Dim oFileWriter As CFileWriter
    Dim strViewTOC As String
    
    Set oFileWriter = New CFileWriter
    
    frmStartProcessing.pbCurrentObject.Min = 1
    frmStartProcessing.pbCurrentObject.Max = oDatabase.Views.Count + 1
    frmStartProcessing.pbCurrentObject.Value = 1
    
    For Each oView In oDatabase.Views
        If oView.SystemObject = False Then
            With oFileWriter
                .filename = "View." & oView.owner & "." & oView.Name & ".htm"
                .Path = frmSelectDirectory.txtDirectory
                .FileData = CreateViewHTMLDoc(oView)
                .WriteToFile
            
                strViewTOC = strViewTOC & TOCViewD
                strViewTOC = Replace(strViewTOC, "<%VIEWNAME%>", oView.Name)
                strViewTOC = Replace(strViewTOC, "<%VIEWLINK%>", .filename)
            End With
        End If
        frmStartProcessing.pbCurrentObject.Value = frmStartProcessing.pbCurrentObject.Value + 1
        DoEvents
    Next
    strIndex = strIndex & strViewTOC
    TOC = TOC & strViewTOC
End Function

Private Function CreateViewHTMLDoc(ByRef dbObject As View) As String
    Dim strHTML As String
    Dim oView As New CView
    
    With oView
        Set .ViewObject = dbObject
        strHTML = .Header
        strHTML = strHTML & .Properties
        strHTML = strHTML & .Column
        strHTML = strHTML & .Source
    End With
    
    strHTML = strHTML & EndHead
    CreateViewHTMLDoc = strHTML
End Function

Public Function StoredProcedureDOC()
    Dim oStoredProcedure As StoredProcedure
    Dim oFileWriter As CFileWriter
    Dim strStoredProcedureTOC As String
    
    Set oFileWriter = New CFileWriter
    
    frmStartProcessing.pbCurrentObject.Min = 1
    frmStartProcessing.pbCurrentObject.Max = oDatabase.StoredProcedures.Count + 1
    frmStartProcessing.pbCurrentObject.Value = 1
    For Each oStoredProcedure In oDatabase.StoredProcedures
        If oStoredProcedure.SystemObject = False Then
            With oFileWriter
                .filename = "StoredProcedure." & oStoredProcedure.owner & "." & oStoredProcedure.Name & ".htm"
                .Path = frmSelectDirectory.txtDirectory
                .FileData = CreateStoredProcedureHTMLDoc(oStoredProcedure)
                .WriteToFile
                
                strStoredProcedureTOC = strStoredProcedureTOC & TOCStoredProcedureD
                strStoredProcedureTOC = Replace(strStoredProcedureTOC, "<%STOREDPROCEDURENAME%>", oStoredProcedure.Name)
                strStoredProcedureTOC = Replace(strStoredProcedureTOC, "<%STOREDPROCEDURELINK%>", .filename)
            End With
        End If
        
        frmStartProcessing.pbCurrentObject.Value = frmStartProcessing.pbCurrentObject.Value + 1
        DoEvents
    Next
    strIndex = strIndex & strStoredProcedureTOC
    TOC = TOC & strStoredProcedureTOC
End Function

Private Function CreateStoredProcedureHTMLDoc(ByRef oStoredProcedure As StoredProcedure) As String
    Dim ObjStoredProcedure As CStoredProcedure
    Dim strHTML As String
            
    Set ObjStoredProcedure = New CStoredProcedure
    With ObjStoredProcedure
        Set .StoredProcedureObject = oStoredProcedure
        strHTML = .Header
        strHTML = strHTML & .Properties
        strHTML = strHTML & .Documentation
        strHTML = strHTML & .Parameters
        strHTML = strHTML & .Dependant
        strHTML = strHTML & .SourceText
        strHTML = strHTML & EndHead
    End With
    CreateStoredProcedureHTMLDoc = strHTML
End Function

Public Function UserDefinedFunctionsDOC()
    Dim oUserDefinedFunction As UserDefinedFunction
    Dim oFileWriter As CFileWriter
    Dim strUserDefinedFunctionTOC As String
 
    Set oFileWriter = New CFileWriter
    
    If oDatabase.UserDefinedFunctions.Count = 0 Then
        Exit Function
    End If
    
    frmStartProcessing.pbCurrentObject.Min = 1
    frmStartProcessing.pbCurrentObject.Max = oDatabase.UserDefinedFunctions.Count + 1
    frmStartProcessing.pbCurrentObject.Value = 1
    For Each oUserDefinedFunction In oDatabase.UserDefinedFunctions
        If oUserDefinedFunction.SystemObject = False Then
            With oFileWriter
                .filename = "UserDefinedFunction." & oUserDefinedFunction.owner & "." & oUserDefinedFunction.Name & ".htm"
                .Path = frmSelectDirectory.txtDirectory
                .FileData = CreateUserDefinedFunctionHTMLDoc(oUserDefinedFunction)
                .WriteToFile
                
                strUserDefinedFunctionTOC = strUserDefinedFunctionTOC & TOCUserDefinedFunctionD
                strUserDefinedFunctionTOC = Replace(strUserDefinedFunctionTOC, "<%USERDEFINEDFUNCTIONNAME%>", oUserDefinedFunction.Name)
                strUserDefinedFunctionTOC = Replace(strUserDefinedFunctionTOC, "<%USERDEFINEDFUNCTIONLINK%>", .filename)
            End With
        End If
        
        frmStartProcessing.pbCurrentObject.Value = frmStartProcessing.pbCurrentObject.Value + 1
        DoEvents
    Next
    strIndex = strIndex & strUserDefinedFunctionTOC
    TOC = TOC & strUserDefinedFunctionTOC
End Function

Private Function CreateUserDefinedFunctionHTMLDoc(ByRef oUserDefinedFunction As UserDefinedFunction) As String
    Dim ObjUserDefinedFunction As CFunctions
    Dim strHTML As String
            
    Set ObjUserDefinedFunction = New CFunctions
    With ObjUserDefinedFunction
        Set .UserDefinedFunctionObject = oUserDefinedFunction
        strHTML = .Header
        strHTML = strHTML & .Properties
        strHTML = strHTML & .Documentation
        strHTML = strHTML & .Parameters
        strHTML = strHTML & .Dependant
        strHTML = strHTML & .SourceText
        strHTML = strHTML & EndHead
    End With
    CreateUserDefinedFunctionHTMLDoc = strHTML
End Function



