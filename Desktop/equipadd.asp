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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_NationStar_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.DesktopInventory (DInType,DInModel,DInSerialNum,DInStatus,DInNotes,DInSite,DInImageVer,DInArcherNum,DInCubeLocation,DInTicketNum,DInAssetTagNum,DInCostCenter,DInAssignedUser,DInUserID,DInDeployedDate,DInEndDate,DInComputerName) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("DInType")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 50, Request.Form("DInModel")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("DInSerialNum")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("DInStatus")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 1000, Request.Form("DInNotes")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, 50, Request.Form("DInSite")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, 50, Request.Form("DInImageVer")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 201, 1, 50, Request.Form("DInArcherNum")) ' adLongVarChar 
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 201, 1, 50, Request.Form("DInCubeLocation")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 201, 1, 50, Request.Form("DInTicketNum")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 5, 1, -1, MM_IIF (Request.Form("DInAssetTagNum"), Request.Form("DInAssetTagNum"), null)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 5, 1, -1, MM_IIF (Request.Form("DInCostCenter"), Request.Form("DInCostCenter"), null)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 201, 1, 50, Request.Form("DInAssignedUser")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 201, 1, 50, Request.Form("DInUserID")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 135, 1, 10, MM_IIF(Request.Form("DInDeployedDate"), Request.Form("DInDeployedDate"), null)) ' adDBTimeStamp
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 135, 1, 10, MM_IIF(Request.Form("DInEndDate"), Request.Form("DInEndDate"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 201, 1, 50, Request.Form("DInComputerName")) ' adLongVarChar	
	MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "equip_search.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>

<%
Dim DIntype
Dim DIntype_cmd
Dim DIntype_numRows

Set DIntype_cmd = Server.CreateObject ("ADODB.Command")
DIntype_cmd.ActiveConnection = MM_NationStar_STRING
DIntype_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '1'  ORDER BY equipment ASC" 
DIntype_cmd.Prepared = true

Set DIntype = DIntype_cmd.Execute
DIntype_numRows = 0
%>


<%
Dim DInmodel
Dim DInmodel_cmd
Dim DInmodel_numRows

Set DInmodel_cmd = Server.CreateObject ("ADODB.Command")
DInmodel_cmd.ActiveConnection = MM_NationStar_STRING
DInmodel_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '2'  ORDER BY equipment ASC" 
DInmodel_cmd.Prepared = true

Set DInmodel = DInmodel_cmd.Execute
DInmodel_numRows = 0
%>

<%
Dim DInimagever
Dim DInimagever_cmd
Dim DInimagever_numRows

Set DInimagever_cmd = Server.CreateObject ("ADODB.Command")
DInimagever_cmd.ActiveConnection = MM_NationStar_STRING
DInimagever_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '3'  ORDER BY equipment ASC" 
DInimagever_cmd.Prepared = true

Set DInimagever = DInimagever_cmd.Execute
imagever_numRows = 0
%>

<%
Dim DInstatus
Dim DInstatus_cmd
Dim DInstatus_numRows

Set DInstatus_cmd = Server.CreateObject ("ADODB.Command")
DInstatus_cmd.ActiveConnection = MM_NationStar_STRING
DInstatus_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '5'  ORDER BY equipment ASC" 
DInstatus_cmd.Prepared = true

Set DInstatus = DInstatus_cmd.Execute
DInstatus_numRows = 0
%>

<%
Dim DInsite
Dim DInsite_cmd
Dim DInsite_numRows

Set DInsite_cmd = Server.CreateObject ("ADODB.Command")
DInsite_cmd.ActiveConnection = MM_NationStar_STRING
DInsite_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '4'  ORDER BY equipment ASC" 
DInsite_cmd.Prepared = true

Set DInsite = DInsite_cmd.Execute
DInsite_numRows = 0
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
			<h1>Desktop Equipment Asset Editor</h1>
        </div>
        </div>
        <div id="navigationwrap">
        <div id="navigation">
			<ul id="menu-bar">
				<li><a href="/desktop/">Dashboard</a></li>
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
        <div id="content" style="height:400px;" >
		
			
			
<form ACTION="<%=MM_editAction%>" method="POST" name="form1" id="form1">
	<table width="514" border="0">
		<tr><td width="53">Type: <br>
			<select name="DIntype" id="DIntype">
				<%
				While (NOT DIntype.EOF)
				%>
				<option value="<%=(DInType.Fields.Item("equipment").Value)%>"><%=(DInType.Fields.Item("equipment").Value)%></option>
				<%
				DIntype.MoveNext()
				Wend
				If (DIntype.CursorType > 0) Then
				DIntype.MoveFirst
				Else
				DIntype.Requery
				End If
				%>
			</select></td>

			<td width="122">Model: <br>
			<label for="DInmodel"></label>
			<select name="DInmodel" id="DInmodel">
				<%
				While (NOT DInmodel.EOF)
				%>
				<option value="<%=(DInModel.Fields.Item("equipment").Value)%>"><%=(DInModel.Fields.Item("equipment").Value)%></option>
				<%
				DInmodel.MoveNext()
				Wend
				If (DInmodel.CursorType > 0) Then
				DInmodel.MoveFirst
				Else
				DInmodel.Requery
				End If
				%>
				</select></td>

			<td width="105">Image Version: <br>
			<label for="DInimagever"></label>
				<select name="DInimagever" id="DInimagever">
					<%
					While (NOT DInimagever.EOF)
					%>
					<option value="<%=(DInImageVer.Fields.Item("equipment").Value)%>"><%=(DInImageVer.Fields.Item("equipment").Value)%></option>
					<%
					DInimagever.MoveNext()
					Wend
					If (DInimagever.CursorType > 0) Then
					DInimagever.MoveFirst
					Else
					DInimagever.Requery
					End If
					%> 
				</select></td>      


			<td width="95">Status:<br>
			<label for="DInstatus"></label>
				<select name="DInstatus" id="DInstatus">
					<%
					While (NOT DInstatus.EOF)
					%>
					<option value="<%=(DInStatus.Fields.Item("equipment").Value)%>"><%=(DInStatus.Fields.Item("equipment").Value)%></option>
					<%
					DInstatus.MoveNext()
					Wend
					If (DInstatus.CursorType > 0) Then
					DInstatus.MoveFirst
					Else
					DInstatus.Requery
					End If
					%> 
				</select></td>

			<td width="117">Location:<br>
			<label for="DInsite"></label>
				<select name="DInsite" id="DInsite">
					<%
					While (NOT DInsite.EOF)
					%>
					<option value="<%=(DInSite.Fields.Item("equipment").Value)%>"><%=(DInSite.Fields.Item("equipment").Value)%></option>
					<%
					DInsite.MoveNext()
					Wend
					If (DInsite.CursorType > 0) Then
					DInsite.MoveFirst
					Else
					DInsite.Requery
					End If
					%> 
				</select></td>
					
			<td>Computer Name: <br>
			<label for="DInComputerName"></label>
				<input name="DInComputerName" type="text" id="DInComputerName" maxlength="40" class="txtBox"/></td>
		</tr>
		
		<tr><td>&nbsp;</td>
			<td>Serial: <br>
				<label for="DInserialNum"></label>
					<input name="DInserialNum" type="text" id="DInserialNum" maxlength="50" class="txtBox"/> 
			</td>
			<td>Asset Tag: <br>
				<label for="DInAssetTagNum"></label>
				<input name="DInAssetTagNum" type="text" id="DInAssetTagNum" maxlength="40" class="txtBox"/>
			</td>
			<td>Cube #: <br>
				<label for="DInCubeLocation"></label>
				<input name="DInCubeLocation" type="text" id="DInCubeLocation" maxlength="40" class="txtBox"/>
			</td>
		</tr>
	</table>

<hr>
<div style="float:left; margin: 5px;">
	Ticket #: <br />
	<label for="DInticketnum"></label>
	<input name="DInticketnum" type="text" id="DInticketnum" maxlength="20" class="txtBox"/>
</div>
<div style="float:left; margin: 5px;">
	Assigned: <br />
	<label for="DInassigneduser"></label>
	<input name="DInassigneduser" type="text" id="DInassigneduser" maxlength="40" class="txtBox"/> 
</div>
<div style="float:left; margin: 5px;">	
	User ID (AD ID): <br />
	<label for="DInUserID"></label>
	<input name="DInUserID" type="text" id="DInUserID" maxlength="40" class="txtBox"/>
</div>
<div style="float: left; margin: 5px;">	
	Cost Center: <br />
	<label for="DInCostCenter"></label>
	<input name="DInCostCenter" type="text" id="DInCostCenter" maxlength="40" class="txtBox"/>
</div>
<div style="float: left; margin: 5px;">	
	Archer Number: <br />
	<label for="DInArcherNum"></label>
	<input name="DInArcherNum" type="text" id="DInArcherNum" maxlength="40" class="txtBox"/>
</div>
<div style="float:left; margin: 5px;">	
	Deployed: <br />
	<label for="DIndeployeddate"></label>
	<input name="DIndeployeddate" type="text" id="DIndeployeddate" maxlength="40" class="txtBox"/> 
</div>
<div style="float:left; margin: 5px;">	
	Returned Date: <br />
	<label for="DInenddate"></label>
	<input name="DInenddate" type="text" id="DInenddate" maxlength="40" class="txtBox"/> 
</div>
<div style="float:left; margin: 5px;">	
	Notes: <br />
	<label for="DInnotes"></label>
	<textarea name="DInnotes" cols="60" rows="4" type="text" id="DInnotes" maxlength="1500"></textarea><br /><br />
	<input name="submit" type="submit" value="     Add     " class="btn"/>
</div>
	<input type="hidden" name="MM_insert" value="form1" />
	
</form>			
			
			
		
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
DInmodel.Close()
Set DInmodel = Nothing
%>

<%
DInimagever.Close()
Set DInimagever = Nothing
%>

<%
DInstatus.Close()
Set DInstatus = Nothing
%>

<%
DInsite.Close()
Set DInsite = Nothing
%>




















