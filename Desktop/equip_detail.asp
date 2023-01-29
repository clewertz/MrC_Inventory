<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/NationStar.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>

<%
Dim serial_num
serial_num = ""
If (Request.QueryString("DInSerialNum") <> "") Then 
  serial_num = Request.QueryString("DInSerialNum")
End If
%>

<%
Dim equip_edit
Dim equip_edit_cmd
Dim equip_edit_numRows

Set equip_edit_cmd = Server.CreateObject ("ADODB.Command")
equip_edit_cmd.ActiveConnection = MM_NationStar_STRING
equip_edit_cmd.CommandText = "SELECT * FROM dbo.DesktopInventory WHERE DInSerialNum = ?" 
equip_edit_cmd.Prepared = true
equip_edit_cmd.Parameters.Append equip_edit_cmd.CreateParameter("param1", 200, 1, 255, serial_num) ' adVarChar

Set equip_edit = equip_edit_cmd.Execute
equip_edit_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title>Nationstar Desktop Dashboard</title>
<!--[if IE]><script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script><![endif]-->
<link rel="stylesheet" type="text/css" href="dash.css" />

</head>
<body>
    <div id="wrapper">
        <div id="headerwrap">
        <div id="header">
            <img src="graphics/NSM-logo2.png" style="padding: 8px;">
			<h1>Desktop Equipment Details</h1>
        </div>
        </div>
        <div id="navigationwrap">
        <div id="navigation">
			<ul id="menu-bar">
				<li><a href="/nationstar/desktop/">Dashboard</a></li>
				<li class="active"><a href="equip_search.asp">Equipment</a></li>
				<li><a href="#">Administrative</a>
					<ul>
					<li><a href="https://wd5.myworkday.com/nationstar/login.flex">Workday</a></li>
					<li><a href="https://login.daptiv.com/">Daptiv</a></li>
					<li><a href="http://teams/sites/desktopservices/">Sharepoint</a></li>
					</ul>
				</li>
			</ul>
        </div>
        </div>
        <div id="contentwrap">
        <div id="content" >
		
		<table class="eqTable">
			<tr><td><h3>Type: </h3>
			  <%=(equip_edit.Fields.Item("DInType").Value)%></td>
				<td><h3>Model:</h3>
			  <%=(equip_edit.Fields.Item("DInModel").Value)%></td>
				<td><h3>Location:</h3>
			  <%=(equip_edit.Fields.Item("DInSite").Value)%></td>
				<td><h3>Image:</h3>
			  <%=(equip_edit.Fields.Item("DInImageVer").Value)%></td>     
				<td><h3>Status:</h3>
			  <%=(equip_edit.Fields.Item("DInStatus").Value)%></td></tr>
				
			<tr><td><h3>Ticket #:</h3>
			  <%=(equip_edit.Fields.Item("DInTicketNum").Value)%></td>
				<td><h3>Serial:</h3>
			  <%=(equip_edit.Fields.Item("DInSerialNum").Value)%></td>
			  
				<td><h3>Assigned:</h3>
			  <%=(equip_edit.Fields.Item("DInAssignedUser").Value)%></td>
				<td><h3>Deployed:</h3>
			  <%=(equip_edit.Fields.Item("DInDeployedDate").Value)%></td>
				<td><h3>End Date:</h3>
			  <%=(equip_edit.Fields.Item("DInEndDate").Value)%></td></tr>

			<tr><td colspan="5"><h3>Notes:</h3>
			  <%=(equip_edit.Fields.Item("DInNotes").Value)%></td></tr>

		</table>		
			
        </div>
        </div>
        <div id="footerwrap">
        <div id="footer">
            <p>&nbsp;</p>
        </div>
        </div>
    </div>
</body>
</html>


<%
equip_edit.Close()
Set equip_edit = Nothing
%>
