Attribute VB_Name = "HTMLConst"
Option Explicit


Public Const CSSStyleSheet = "Style.css"

Public Const RowColour_1 = "RowColour_1"
Public Const RowColour_2 = "RowColour_2"

'------------------------------------------------------------------------------------------------------------------------
Public Const HHeader1 = "<TABLE style='WIDTH: 100%' cellSpacing='0' cellPadding='0' border='0'><TR><TD><H4>Welcome to Indus SQL Server document builder</H4></TD></TR><TR><TD><BLOCKQUOTE dir='ltr' style='MARGIN-RIGHT: 0px'><P>This wizard will guide you to document your SQL Server eaisily &amp; efficiently</P></BLOCKQUOTE></TD></TR></TABLE>"
Public Const HHeader2 = "<TABLE style='WIDTH: 100%' cellSpacing='0' cellPadding='0' border='0'><TR><TD><H5>To connect to Microsoft SQL Server, you must specify the server, user name and password.</H5></TD></TR></TABLE>"
Public Const HHeader3 = "<TABLE style='WIDTH: 100%' cellSpacing='0' cellPadding='0' border='0'><TR><TD><H5><UL><LI>Enter output directory in the box below or use browse for folder button to select the output directory.</LI><LI>Choose HTML help compiler options.</LI><LI>Choose database objects for which you want to generate documents.</LI></UL></H5></TD></TR></TABLE>"
Public Const HHeader4 = "<TABLE style='WIDTH: 100%' cellSpacing='0' cellPadding='0' border='0'><TR><TD><H5>Start Building Document...</H5></TD></TR></TABLE>"

'------------------------------------------------------------------------------------------------------------------------
Public Const StartTitle = "<TITLE>"
Public Const StartDIV = "<div class=H1><%HEADING%></div>"
Public Const EndTitle = "</TITLE>"

'------------------------------------------------------------------------------------------------------------------------
Public Const StartHead = "<HTML><HEAD><link rel='stylesheet' type='text/css' href='" & CSSStyleSheet & "'></HEAD><BODY>"
Public Const EndHead = "</BODY></HTML>"

'------------------------------------------------------------------------------------------------------------------------
Public Const PropDBHeadHTML = "<DIV class='h2'>Server Name: <%SERVERNAME%></DIV>" & _
    "<DIV class='h2'>Database Name: <%DATABASENAME%></DIV>" & _
    "<DIV class='h2'>Database System: <%DATABASE_SYSTEM%></DIV>"

'------------------------------------------------------------------------------------------------------------------------
Public Const PropTableHeadHTML = "<id=Table Properties><div class=h3>Table Properties</div>" & vbCrLf & _
    "<TABLE border=0 cellPadding=2 cellSpacing=1 width=100%>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD width=30%>Property Name</TD><TD width=70%>Property Value</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD>Creation Date</TD><TD><%CREATION_DATE%></TD></TR>" & vbCrLf & _
    "<TR class=" & RowColour_2 & "><TD>Data Space Used</TD><TD><%DATA_SPACE_USED%></TD></TR>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD>Number of Rows</TD><TD><%NO_OF_ROWS%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE></id=Table Properties>"
Public Const PropTableDescHTML = "<id=Table Description><div class=h3>Table Description</div>" & vbCrLf & _
    "<TABLE border=0 cellPadding=2 cellSpacing=1 width=100%>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD width=30%>Property Name</TD><TD width=70%>Property Value</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD>Description</TD><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE></id=Table Description>"
Public Const PropTableColumnTabStartHTML = "<DIV class=h3>Table Columns</DIV>" & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>"
Public Const PropTableColumnHeadHTML = "<THEAD>" & _
    "<TR>" & _
    "<TD width='15%'>Column Name</TD>" & _
    "<TD width='25%'>Description</TD>" & _
    "<TD width='6%'>In Primary Key</TD>" & _
    "<TD width='7%'>Data Type</TD>" & _
    "<TD width='6%'>Length</TD>" & _
    "<TD width='6%'>Precision</TD>" & _
    "<TD width='6%'>Scale</TD>" & _
    "<TD width='6%'>Allow Nulls</TD>" & _
    "<TD width='7%'>Default</TD>" & _
    "<TD width='7%'>Rule</TD>" & _
    "<TD width='7%'>Identity (ID)</TD>" & _
    "<TR><THEAD>"
Public Const PropTableColumnDetailsHTML = "<TBODY class=small>"
Public Const PropTableColumnTabEndHTML = "</TBODY></TABLE>"
Public Const PropTableIndexStartHTML = "<DIV class=h3>Table Indexes</DIV>" & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>"
Public Const PropTableIndexHeadHTML = "<THEAD>" & _
    "<TR>" & _
    "<TD width='15%'>Index Name</TD>" & _
    "<TD width='25%'>Columns</TD>" & _
    "<TD width='6%'>In Primary Index</TD>" & _
    "<TD width='6%'>Clustered</TD>" & _
    "<TD width='6%'>Unique Index</TD>" & _
    "<TD width='6%'>Size KB</TD>" & _
    "<TR><THEAD>"
Public Const PropTableIndexDetailsHTML = "<TBODY class=small>"
Public Const PropTableIndexEndHTML = "</TBODY></TABLE>"
Public Const PropTableDependantHTML = "<DIV class=h3>Table Dependant Listing</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Object Name</TD>" & vbCrLf & _
    "<TD width='20%'>Object Type</TD>" & vbCrLf & _
    "<TD width='20%'>Dep Level</TD>" & vbCrLf & _
    "</TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf
Public Const PropTableReferencedTablesHTML = "<DIV class=h3>Relationships Constraints (Referenced Tables)</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Table</TD>" & vbCrLf & _
    "<TD width='20%'>Key</TD>" & vbCrLf & _
    "<TD width='20%'>Check existing data on creation</TD>" & vbCrLf & _
    "</TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf
Public Const PropTableReferencingTablesHTML = "<DIV class=h3>Relationships Constraints (Referencing Tables)</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Table</TD>" & vbCrLf & _
    "<TD width='20%'>Key</TD>" & vbCrLf & _
    "<TD width='20%'>Check existing data on creation</TD>" & vbCrLf & _
    "</TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf

Public Const PropAllTableHTML = "<DIV class=H1>Tables</DIV><BR><TABLE cellSpacing=1 cellPadding=2 width=100% border=0><THEAD>" & vbCrLf & _
    "<TR><TD width='30%'>Table Name</TD>" & vbCrLf & _
    "<TD align=right width='15%'>Number of Rows</TD>" & vbCrLf & _
    "<TD align=right width='15%'>Data Size KB</TD>" & vbCrLf & _
    "<TD align=right width='15%'>Index Size KB</TD>" & vbCrLf & _
    "<TD align=right width='10%'>Creation Date</TD></TR></THEAD>"


'------------------------------------------------------------------------------------------------------------------------
Public Const PropDatabaseDetails = "<DIV class='h1'>Database Details</DIV><BR>" & _
    "<DIV class='h3'>Database Properties</DIV>" & _
    "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<THEAD><TR><TD width='30%'>Property Name</TD><TD width='70%'>Property Value</TD></TR></THEAD>" & _
    "<TR class='RowColour_1'><TD>Server Name</TD><TD><%SERVERNAME%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Database Name</TD><TD><%DATABASENAME%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Database System</TD><TD><%DATABASESYSTEM%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Database Version</TD><TD><%DATABASEVERSION%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Run Date</TD><TD><%RUNDATE%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Creation Date</TD><TD><%CREATIONDATE%></TD></TR>" & _
    "</TABLE>"


'------------------------------------------------------------------------------------------------------------------------
Public Const PropTriggerHeadHTML = "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<DIV class='h3'>Trigger Properties</DIV>" & _
    "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<THEAD><TR><TD width='30%'>Property Name</TD><TD width='70%'>Property Value</TD></TR></THEAD>" & _
    "<TBODY><TR class='RowColour_1'><TD>Table Name</TD><TD><A href='<%TABLELINK%>'><%TABLENAME%></A></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Creation Date</TD><TD><%CREATIONDATE%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Precision</TD><TD><%PRECISION%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Insert Trigger</TD><TD><%INSERTTRIGGER%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Update Trigger</TD><TD><%UPDATETRIGGER%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Delete Trigger</TD><TD><%DELETETRIGGER%></TD></TR></TBODY>" & _
    "</TABLE>"
Public Const PropTriggerDescHTML = "<div class='h3'>Trigger Description</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Description</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"
Public Const PropTriggerSourceHTML = "<div class='h3'>Trigger Source</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Source</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"


'------------------------------------------------------------------------------------------------------------------------
Public Const PropViewHeadHTML = "<id=View Properties><div class=h3>View Properties</div>" & vbCrLf & _
    "<Table border=0 cellPadding=2 cellSpacing=1 width=100%>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD width=30%>Property Name</TD><TD width=70%>Property Value</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD>Creation Date</TD><TD><%CREATION_DATE%></TD></TR>" & vbCrLf & _
    "<TR class=" & RowColour_2 & "><TD>Is Schema Bound</TD><TD><%IS_SCHEMA_BOUND%></TD></TR>" & vbCrLf & _
    "</TBODY></table></id=View Properties>"
Public Const PropViewDescHTML = "<id=View Description><div class=h3>View Description</div>" & vbCrLf & _
    "<table border=0 cellPadding=2 cellSpacing=1 width=100%>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD width=30%>Property Name</TD><TD width=70%>Property Value</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD>Description</TD><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></Table></id=View Description>"
Public Const PropViewColumnTabStartHTML = "<DIV class=h3>View Columns</DIV>" & _
    "<Table cellSpacing=1 cellPadding=2 width='100%' border=0>"
Public Const PropViewColumnHeadHTML = "<THEAD>" & _
    "<TR>" & _
    "<TD width='15%'>Column Name</TD>" & _
    "<TD width='15%'>Data Type</TD>" & _
    "<TD width='10%'>Length</TD>" & _
    "<TD width='10%'>Precision</TD>" & _
    "<TD width='10%'>Scale</TD>" & _
    "<TR><THEAD>"
Public Const PropViewColumnDetailsHTML = "<TBODY class=small>"
Public Const PropViewColumnTabEndHTML = "</TBODY></Table>"
Public Const PropViewSourceHTML = "<id=View Source><div class=h3>View Source</div>" & vbCrLf & _
    "<table border=0 cellPadding=2 cellSpacing=1 width=100%>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD width=30%>SQL Script to rebuild the view</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class=" & RowColour_1 & "><TD><%SOURCE%></TD></TR>" & vbCrLf & _
    "</TBODY></Table></id=View Source>"


'------------------------------------------------------------------------------------------------------------------------
Public Const PropStoredProcedureHeadHTML = "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<DIV class='h3'>Stored Procedure Properties</DIV>" & _
    "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<THEAD><TR><TD width='30%'>Property Name</TD><TD width='70%'>Property Value</TD></TR></THEAD>" & _
    "<TBODY><TR class='RowColour_1'><TD>Creation Date</TD><TD><%CREATIONDATE%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Encrypted</TD><TD><%ENCRYPTED%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Start Up Stored Procedure</TD><TD><%STARTUPSTOREDPROCEDURE%></TD></TR></TBODY>" & _
    "</TABLE>"
Public Const PropStoredProcedureDescHTML = "<div class='h3'>Stored Procedure Description</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Description</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"
Public Const PropStoredProcedureSourceHTML = "<div class='h3'>Stored Procedure Source</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Source</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"

Public Const PropStoredProcedureParameterHTML = "<DIV class=h3>Stored Procedure Parameters</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Parameter Name</TD>" & vbCrLf & _
    "<TD width='20%'>Data Type</TD>" & vbCrLf & _
    "<TD width='20%'>Length</TD>" & vbCrLf & _
    "<TD width='20%'>Output Parameter</TD></TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf

Public Const PropStoredProcedureDependantHTML = "<DIV class=h3>Stored Procedure Dependant Listing</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Object Name</TD>" & vbCrLf & _
    "<TD width='20%'>Object Type</TD>" & vbCrLf & _
    "<TD width='20%'>Dep Level</TD>" & vbCrLf & _
    "</TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf
'----------------------------------------------------------------------------------------------------------------


Public Const PropUserDefinedFunctionHeadHTML = "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<DIV class='h3'>User Defined Function Properties</DIV>" & _
    "<TABLE cellSpacing='1' cellPadding='2' width='100%' border='0'>" & _
    "<THEAD><TR><TD width='30%'>Property Name</TD><TD width='70%'>Property Value</TD></TR></THEAD>" & _
    "<TBODY><TR class='RowColour_1'><TD>Creation Date</TD><TD><%CREATIONDATE%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Encrypted</TD><TD><%ENCRYPTED%></TD></TR>" & _
    "<TR class='RowColour_1'><TD>Is Deterministic</TD><TD><%ISDETERMINISTIC%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Is Schema Bound</TD><TD><%ISSCHEMABOUND%></TD></TR>" & _
    "<TR class='RowColour_2'><TD>Type</TD><TD><%TYPE%></TD></TR>" & _
    "</TBODY></TABLE>"
Public Const PropUserDefinedFunctionDescHTML = "<div class='h3'>User Defined Function Description</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Description</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"
Public Const PropUserDefinedFunctionSourceHTML = "<div class='h3'>User Defined Function Source</div>" & vbCrLf & _
    "<TABLE border='0' cellPadding='2' cellSpacing='1' width='100%'>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR><TD 'width=100%'>Source</TD></TR>" & vbCrLf & _
    "</THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf & _
    "<TR class='" & RowColour_1 & "'><TD><%DESCRIPTION%></TD></TR>" & vbCrLf & _
    "</TBODY></TABLE>"

Public Const PropUserDefinedFunctionParameterHTML = "<DIV class=h3>User Defined Function Parameters</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Parameter Name</TD>" & vbCrLf & _
    "<TD width='20%'>Data Type</TD>" & vbCrLf & _
    "<TD width='20%'>Length</TD>" & vbCrLf & _
    "<TD width='20%'>Parameter Direction</TD></TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf

Public Const PropUserDefinedFunctionDependantHTML = "<DIV class=h3>User Defined Function Dependant Listing</DIV>" & vbCrLf & _
    "<TABLE cellSpacing=1 cellPadding=2 width='100%' border=0>" & vbCrLf & _
    "<THEAD>" & vbCrLf & _
    "<TR>" & vbCrLf & _
    "<TD width='40%'>Object Name</TD>" & vbCrLf & _
    "<TD width='20%'>Object Type</TD>" & vbCrLf & _
    "<TD width='20%'>Dep Level</TD>" & vbCrLf & _
    "</TR></THEAD>" & vbCrLf & _
    "<TBODY>" & vbCrLf

