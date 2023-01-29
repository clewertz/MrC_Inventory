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
<link rel="stylesheet" type="text/css" href="desktop.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Equip Detail</title>
</head>

<body>

<div id = "header">Equipment Detail</div>
<div id="menu" align="center"></div>
<br>

<table width="800" border="0">
  <tr>
    <td width="244">  
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" --></td>
    <td width="546"><form ACTION="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="514" border="1">
  <tr align="center">
    <td width="150"><b>Type: </b><br><br>
      <%=(equip_edit.Fields.Item("DInType").Value)%>
	</td>
    
	<td width="150"><b>Model:</b><br><br>
      <%=(equip_edit.Fields.Item("DInModel").Value)%>
	</td>
	  
	  <td width="150"><b>Location:</b><br><br>
      <%=(equip_edit.Fields.Item("DInSite").Value)%>
	</td>
        </select></td>
        
  <td width="150"><b>Image Vers:</b><br><br>
      <%=(equip_edit.Fields.Item("DInImageVer").Value)%>
	</td>     


	<td width="150"><b>Status:</b><br><br>
      <%=(equip_edit.Fields.Item("DInStatus").Value)%>
	</td>
	  
  </tr>
  <tr align="center">
    

	<td><b>Ticket #:</b><br><br>
      <%=(equip_edit.Fields.Item("DInTicketNum").Value)%>
	</td>


    <td><b>Serial:</b><br><br>
       <%=(equip_edit.Fields.Item("DInSerialNum").Value)%>
	</td>
	  
	  
    <td><b>Assigned:</b><br><br>
      <%=(equip_edit.Fields.Item("DInAssignedUser").Value)%>
	</td>
    
    <td><b>Deployed:</b><br><br>
      <%=(equip_edit.Fields.Item("DInDeployedDate").Value)%>
	</td>
    
    <td><b>End Date:</b><br><br>
      <%=(equip_edit.Fields.Item("DInEndDate").Value)%>
	</td>
    
	
  </tr>
  <tr align="center">
    
	
    <td colspan="5"><b>Notes:</b><br><br>
      <%=(equip_edit.Fields.Item("DInNotes").Value)%>
	</td>
    </tr>

</table>
<br><br>
<img src="graphics/bar_break.png" width="800" height="10" />

<div id="mainbody">

</body>
</html>

<%
equip_edit.Close()
Set equip_edit = Nothing
%>
