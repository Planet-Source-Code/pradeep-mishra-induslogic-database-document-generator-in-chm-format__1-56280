Attribute VB_Name = "SQLServer"
Option Explicit

Public mServerApp As SQLDMO.Application     'SQL Server Application Object
Public mServer As SQLDMO.SQLServer2         'SQL Server Database Object
Public mIsConnected As Boolean              'Are you connected to Some Server ?
Public thisServer As String                 'Current Server to which you are connected
Public thisDB As String                     'Current Database
Public oDatabase As Database2               'Curent Database Object


Public FolderPath As String
Public DatabaseIndex As Long

Public Function ConnectToServer(ByVal mAddress As String, ByVal mUName As String, ByVal mPWD As String, ByVal isNTAuth As Boolean, ByRef strErrorMessage As String) As Boolean
    Set mServer = New SQLDMO.SQLServer
    On Error GoTo ConnectError
    Screen.MousePointer = vbHourglass
    With mServer
        If isNTAuth Then
            'If Use NT Authenticaion then
            .LoginSecure = True
            .Connect mAddress
        Else
            .Connect mAddress, mUName, mPWD
        End If
    End With
    
    Screen.MousePointer = vbNormal
    ConnectToServer = True
    thisServer = mAddress
    
    mIsConnected = True
    Exit Function
ConnectError:
    strErrorMessage = Err.Description
    ConnectToServer = False
    Screen.MousePointer = vbNormal
End Function

Public Function GetTableDescription(ByRef tbTable As Table) As String
    Dim qrDescription As QueryResults
    Dim strSQL As String
    
    strSQL = "SELECT * FROM ::fn_listextendedproperty(NULL, 'user', '" & tbTable.owner & "','table','" & tbTable.Name & "', default, default)"
    Set qrDescription = oDatabase.ExecuteWithResults(strSQL)

    If qrDescription.rows > 0 Then
        GetTableDescription = qrDescription.GetColumnString(1, 4)
    Else
        GetTableDescription = ""
    End If
    
    Set qrDescription = Nothing
End Function

Public Function GetColumnDescription(ByRef tbTable As Table, ByVal strColumnName As String) As String
    Dim qrDescription As QueryResults
    Dim strSQL As String

    strSQL = "SELECT * FROM ::fn_listextendedproperty(NULL, 'user', '" & tbTable.owner & "','table','" & tbTable.Name & "', 'column', '" & strColumnName & "')"
    Set qrDescription = oDatabase.ExecuteWithResults(strSQL)

    If qrDescription.rows > 0 Then
        GetColumnDescription = qrDescription.GetColumnString(1, 4)
    Else
        GetColumnDescription = ""
    End If
    
    Set qrDescription = Nothing
End Function

Public Function ListAllServers(Optional cmbCombo As ComboBox) As NameList
    On Error GoTo ServerListError
    
    Set mServerApp = New SQLDMO.Application
    Dim mNames As SQLDMO.NameList
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Set mNames = mServerApp.ListAvailableSQLServers
    DoEvents
    
    Set ListAllServers = mNames
    
    If Not cmbCombo Is Nothing Then
        cmbCombo.Clear
        Dim temp As Integer
        For temp = 1 To mNames.Count
            cmbCombo.AddItem mNames(temp)
        Next
    End If
    
ServerListError:
    Screen.MousePointer = vbNormal
    Exit Function
End Function

Public Function ListIndexedColumns(idxIndex As index)
    Dim xCtr As Long
    Dim str As String
    Dim idx As Object
    
    For xCtr = 1 To idxIndex.ListIndexedColumns.Count
        Set idx = idxIndex.ListIndexedColumns(xCtr)
        str = str & idx.Name & ","
    Next
    
    str = Left(str, Len(str) - 1)
    
    ListIndexedColumns = str
End Function

Public Function PrimaryKey(tbTable As Table) As String
    If Not tbTable.PrimaryKey Is Nothing Then
        PrimaryKey = tbTable.PrimaryKey.Name
    Else
        PrimaryKey = ""
    End If
End Function

Public Function ClusteredIndex(tbTable As Table) As String
    If Not tbTable.ClusteredIndex Is Nothing Then
        ClusteredIndex = tbTable.ClusteredIndex.Name
    Else
        ClusteredIndex = ""
    End If
End Function

Public Function GetTriggerDescription(ByRef oTrigger As Trigger) As String
    Dim str As String
    str = oTrigger.Text
    str = Replace(FormatText(str), vbCrLf, "<br>")
    GetTriggerDescription = str
End Function

Public Function FormatText(strText As String) As String
    Dim lnStart As Long
    Dim lnEnd As Long
    Dim str As String
    
    str = strText
    
    lnStart = InStr(str, "/*")
    lnEnd = InStr(str, "*/") + 1
    
    If lnEnd > lnStart And lnStart > 1 Then
        str = Mid(str, lnStart, lnEnd)
        
        lnStart = InStr(str, "/*")
        lnEnd = InStr(str, "*/") + 1
        str = Mid(str, lnStart, lnEnd)
        
        FormatText = str
    Else
        FormatText = "Description not available, description is the first comment (lines between /* and */) in the source, please put your descriptive comments."
    End If
    
End Function

Public Function GetStoredProcedureDescription(ByRef oStoredProcedure As StoredProcedure) As String
    Dim str As String
    str = oStoredProcedure.Text
    str = Replace(FormatText(str), vbCrLf, "<br>")
    GetStoredProcedureDescription = str
End Function

Public Function GetUserDefinedFunctionDescription(ByRef oUserDefinedFunction As UserDefinedFunction) As String
    Dim str As String
    str = oUserDefinedFunction.Text
    str = Replace(FormatText(str), vbCrLf, "<br>")
    GetUserDefinedFunctionDescription = str
End Function

Function GetObjDesc(ObjType As SQLDMO_OBJECT_TYPE) As String
    Select Case ObjType
        Case SQLDMOObj_View:
            GetObjDesc = "View"
        Case SQLDMOObj_Alert:
            GetObjDesc = "Alert"
        Case SQLDMOObj_AlertSystem:
            GetObjDesc = "Alert System"
        Case SQLDMOObj_AllButSystemObjects:
            GetObjDesc = "All ButSystem Objects"
        Case SQLDMOObj_AllDatabaseObjects:
            GetObjDesc = "All Database Objects"
        Case SQLDMOObj_AllDatabaseUserObjects:
            GetObjDesc = "All Database User Objects"
        Case SQLDMOObj_Application:
            GetObjDesc = "Application"
        Case SQLDMOObj_AutoProperty:
            GetObjDesc = "Auto Property"
        Case SQLDMOObj_Backup:
            GetObjDesc = "Backup"
        Case SQLDMOObj_BackupDevice:
            GetObjDesc = "Backup Device"
        Case SQLDMOObj_BulkCopy:
            GetObjDesc = "Bulk Copy"
        Case SQLDMOObj_Category:
            GetObjDesc = "Category"
        Case SQLDMOObj_Check:
            GetObjDesc = "Check"
        Case SQLDMOObj_Column:
            GetObjDesc = "Column"
        Case SQLDMOObj_Configuration:
            GetObjDesc = "Configuration"
        Case SQLDMOObj_ConfigValue:
            GetObjDesc = "Config Value"
        Case SQLDMOObj_Database:
            GetObjDesc = "Database"
        Case SQLDMOObj_DatabaseRole:
            GetObjDesc = "Database Role"
        Case SQLDMOObj_DBFile:
            GetObjDesc = "DB File"
        Case SQLDMOObj_DBObject:
            GetObjDesc = "DB Object"
        Case SQLDMOObj_DBOption:
            GetObjDesc = "DB Option"
        Case SQLDMOObj_Default:
            GetObjDesc = "Default"
        Case SQLDMOObj_DistributionArticle:
            GetObjDesc = "Distribution Article"
        Case SQLDMOObj_DistributionDatabase:
            GetObjDesc = "Distribution Database"
        Case SQLDMOObj_DistributionPublication:
            GetObjDesc = "Distribution Publication"
        Case SQLDMOObj_DistributionPublisher:
            GetObjDesc = "Distribution Publisher"
        Case SQLDMOObj_DistributionSubscription:
            GetObjDesc = "Distribution Subscription"
        Case SQLDMOObj_Distributor:
            GetObjDesc = "Distributor"
        Case SQLDMOObj_DRIDefault:
            GetObjDesc = "DRI Default"
        Case SQLDMOObj_FileGroup:
            GetObjDesc = "File Group"
        Case SQLDMOObj_FullTextCatalog:
            GetObjDesc = "Full Text Catalog"
        Case SQLDMOObj_FullTextService:
            GetObjDesc = "Full Text Service"
        Case SQLDMOObj_Group:
            GetObjDesc = "Group"
        Case SQLDMOObj_Index:
            GetObjDesc = "Index"
        Case SQLDMOObj_IntegratedSecurity:
            GetObjDesc = "Integrated Security"
        Case SQLDMOObj_Job:
            GetObjDesc = "Job"
        Case SQLDMOObj_JobFilter:
            GetObjDesc = "Job Filter"
        Case SQLDMOObj_JobHistoryFilter:
            GetObjDesc = "Job History Filter"
        Case SQLDMOObj_JobSchedule:
            GetObjDesc = "Job Schedule"
        Case SQLDMOObj_JobServer:
            GetObjDesc = "Job Server"
        Case SQLDMOObj_JobStep:
            GetObjDesc = "Job Step"
        Case SQLDMOObj_Key:
            GetObjDesc = "Key"
        Case SQLDMOObj_Language:
            GetObjDesc = "Language"
        Case SQLDMOObj_Last:
            GetObjDesc = "Last"
        Case SQLDMOObj_LinkedServer:
            GetObjDesc = "Linked Server"
        Case SQLDMOObj_LinkedServerLogin:
            GetObjDesc = "Linked Server Login"
        Case SQLDMOObj_LogFile:
            GetObjDesc = "Log File"
        Case SQLDMOObj_Login:
            GetObjDesc = "Login"
        Case SQLDMOObj_MergeArticle:
            GetObjDesc = "MergeArticle"
        Case SQLDMOObj_MergeDynamicSnapshotJob:
            GetObjDesc = "Merge Dynamic Snapshot Job"
        Case SQLDMOObj_MergePublication:
            GetObjDesc = "Merge Publication"
        Case SQLDMOObj_MergePullSubscription:
            GetObjDesc = "Merge Pull Subscription"
        Case SQLDMOObj_MergeSubscription:
            GetObjDesc = "Merge Subscription:"
        Case SQLDMOObj_MergeSubsetFilter:
            GetObjDesc = "Merge Subset Filter"
        Case SQLDMOObj_Operator:
            GetObjDesc = "Operator"
        Case SQLDMOObj_Permission:
            GetObjDesc = "Permission"
        Case SQLDMOObj_ProcedureParameter:
            GetObjDesc = "ProcedureParameter"
        Case SQLDMOObj_Publisher:
            GetObjDesc = "Publisher"
        Case SQLDMOObj_QueryResults:
            GetObjDesc = "QueryResults"
        Case SQLDMOObj_RegisteredServer:
            GetObjDesc = "RegisteredServer"
        Case SQLDMOObj_RegisteredSubscriber:
            GetObjDesc = "RegisteredSubscriber"
        Case SQLDMOObj_Registry:
            GetObjDesc = "Registry"
        Case SQLDMOObj_RemoteLogin:
            GetObjDesc = "RemoteLogin"
        Case SQLDMOObj_RemoteServer:
            GetObjDesc = "RemoteServer"
        Case SQLDMOObj_Replication:
            GetObjDesc = "Replication"
        Case SQLDMOObj_ReplicationDatabase:
            GetObjDesc = "Replication Database"
        Case SQLDMOObj_ReplicationSecurity:
            GetObjDesc = "Replication Security"
        Case SQLDMOObj_ReplicationStoredProcedure:
            GetObjDesc = "Replication Stored Procedure"
        Case SQLDMOObj_ReplicationTable:
            GetObjDesc = "Replication Table"
        Case SQLDMOObj_Restore:
            GetObjDesc = "Restore"
        Case SQLDMOObj_Rule:
            GetObjDesc = "Rule"
        Case SQLDMOObj_Schedule:
            GetObjDesc = "Schedule"
        Case SQLDMOObj_ServerGroup:
            GetObjDesc = "Server Group"
        Case SQLDMOObj_ServerRole:
            GetObjDesc = "ServerRole"
        Case SQLDMOObj_SQLServer:
            GetObjDesc = "SQLServer"
        Case SQLDMOObj_StoredProcedure:
            GetObjDesc = "StoredProcedure"
        Case SQLDMOObj_Subscriber:
            GetObjDesc = "Subscriber"
        Case SQLDMOObj_SystemDatatype:
            GetObjDesc = "System Datatype"
        Case SQLDMOObj_SystemTable:
            GetObjDesc = "System Table"
        Case SQLDMOObj_TargetServer:
            GetObjDesc = "Target Server"
        Case SQLDMOObj_TargetServerGroup:
            GetObjDesc = "Target Server Group"
        Case SQLDMOObj_TransactionLog:
            GetObjDesc = "Transaction Log"
        Case SQLDMOObj_TransArticle:
            GetObjDesc = "Trans Article"
        Case SQLDMOObj_Transfer:
            GetObjDesc = "Transfer"
        Case SQLDMOObj_TransPublication:
            GetObjDesc = "Trans Publication"
        Case SQLDMOObj_TransPullSubscription:
            GetObjDesc = "Trans Pull Subscription"
        Case SQLDMOObj_TransSubscription:
            GetObjDesc = "Trans Subscription"
        Case SQLDMOObj_Trigger:
            GetObjDesc = "Trigger"
        Case SQLDMOObj_Unknown:
            GetObjDesc = "Unknown"
        Case SQLDMOObj_User:
            GetObjDesc = "User"
        Case SQLDMOObj_UserDefinedDatatype:
            GetObjDesc = "Datatype"
        Case SQLDMOObj_UserDefinedFunction:
            GetObjDesc = "Function"
        Case SQLDMOObj_UserTable:
            GetObjDesc = "Table"
        Case SQLDMOObj_View:
            GetObjDesc = "View"
    End Select
End Function

Public Function GetUDFType(ByVal UDFtype As SQLDMO_UDF_TYPE) As String
    Select Case UDFtype
        Case SQLDMO_UDF_TYPE.SQLDMOUDF_Inline
            GetUDFType = "Inline"
        Case SQLDMO_UDF_TYPE.SQLDMOUDF_Scalar
            GetUDFType = "Scalar"
        Case SQLDMO_UDF_TYPE.SQLDMOUDF_Table
            GetUDFType = "Table"
        Case SQLDMO_UDF_TYPE.SQLDMOUDF_Unknown
            GetUDFType = "Unknown"
    End Select
End Function

Public Function ReturnFileName(ObjectType As String, objOwner As String, objName As String) As String
    ReturnFileName = ObjectType & "." & objOwner & "." & objName & ".htm"
End Function
