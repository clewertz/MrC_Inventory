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







<%
If (CStr(Request("MM_update")) = "update") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_NationStar_STRING
    MM_editCmd.CommandText = "UPDATE dbo.DesktopInventory SET DInType = ?, DInModel = ?, DInSerialNum = ?, DInStatus = ?, DInNotes = ?, DInSite = ?, DInImageVer = ?, DInTicketNum = ?, DInAssignedUser = ?, DInDeployedDate = ?, DInEndDate = ?, DInComputerName = ?, DInAssetTagNum = ?, DInCostCenter = ?, DInCubeLocation = ?, DInUserID = ?, DInArcherNum = ?  WHERE DInSerialNum = ?"
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("DIntype")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 50, Request.Form("DInmodel")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("DInserialnum")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("DInstatus")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 1000, Request.Form("DInnotes")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, 50, Request.Form("DInsite")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, 50, Request.Form("DInimagever")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 201, 1, 50, Request.Form("DInticketnum")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 201, 1, 50, Request.Form("DInassigneduser")) ' adLongVarChar	
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 135, 1, 10, MM_IIF(Request.Form("DIndeployeddate"), Request.Form("DIndeployeddate"), null)) ' adDBTimeStamp
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 135, 1, 10, MM_IIF(Request.Form("DInenddate"), Request.Form("DInenddate"), null)) ' adDBTimeStamp
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 201, 1, 50, Request.Form("DInComputerName")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 5, 1, -1, MM_IIF (Request.Form("DInAssetTagNum"), Request.Form("DInAssetTagNum"), null)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 5, 1, -1, MM_IIF (Request.Form("DInCostCenter"), Request.Form("DInCostCenter"), null)) ' adDouble
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param15", 201, 1, 50, Request.Form("DInCubeLocation")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param16", 201, 1, 50, Request.Form("DInUserID")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param17", 201, 1, 50, Request.Form("DInArcherNum")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param18", 200, 1, 255, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "dash.asp"
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
<link rel="stylesheet" type="text/css" href="desktop.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Equip Edit</title>
</head>

<body>

<div id = "header">Equipment Edit</div>
<div id="menu" align="center"></div>
<br>

<table width="800" border="0">
  <tr>
    <td width="244" valign = "top">  
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" -->
	</td>
    <td width="546"><form ACTION="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="514" border="0">
  <tr>
    <td width="53">Type: <br>
            <select name="DIntype" id="DIntype">
        <option value=""></option>
        <%
While (NOT DIntype.EOF)
%>
        <option value="<%=(DIntype.Fields.Item("equipment").Value)%>"<%If (Not isNull((equip_edit.Fields.Item("DInType").Value))) Then If (CStr(DIntype.Fields.Item("equipment").Value) = CStr((equip_edit.Fields.Item("DInType").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(DIntype.Fields.Item("equipment").Value)%></option>
        <%
  DIntype.MoveNext()
Wend
If (DIntype.CursorType > 0) Then
  DIntype.MoveFirst
Else
  DIntype.Requery
End If
%>

       
      </select>      <br></td>
    
	<td width="122">Model: <br>
      <label for="DInmodel"></label>
<select name="DInmodel" id="DInmodel">
	  <option value=""></option>
        <%
While (NOT DInmodel.EOF)
%>
        <option value="<%=(DInmodel.Fields.Item("equipment").Value)%>"<%If (Not isNull((equip_edit.Fields.Item("DInModel").Value))) Then If (CStr(DInmodel.Fields.Item("equipment").Value) = CStr((equip_edit.Fields.Item("DInModel").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(DInmodel.Fields.Item("equipment").Value)%></option>
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
		  <option value=""></option>
	  
        <%
While (NOT DInimagever.EOF)
%>
        <option value="<%=(DInimagever.Fields.Item("equipment").Value)%>"<%If (Not isNull((equip_edit.Fields.Item("DInImageVer").Value))) Then If (CStr(DInimagever.Fields.Item("equipment").Value) = CStr((equip_edit.Fields.Item("DInImageVer").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(DInimagever.Fields.Item("equipment").Value)%></option>
        <%
  DInimagever.MoveNext()
Wend
If (DInimagever.CursorType > 0) Then
  DInimagever.MoveFirst
Else
  DInimagever.Requery
End If
%> 
      </select> </td>      


		  <td width="95">Status:<br>
        <label for="DInstatus"></label>
        <select name="DInstatus" id="DInstatus">
		  <option value=""></option>
	  
        <%
While (NOT DInstatus.EOF)
%>
        <option value="<%=(DInstatus.Fields.Item("equipment").Value)%>"<%If (Not isNull((equip_edit.Fields.Item("DInStatus").Value))) Then If (CStr(DInstatus.Fields.Item("equipment").Value) = CStr((equip_edit.Fields.Item("DInStatus").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(DInstatus.Fields.Item("equipment").Value)%></option>
        <%
  DInstatus.MoveNext()
Wend
If (DInstatus.CursorType > 0) Then
  DInstatus.MoveFirst
Else
  DInstatus.Requery
End If
%> 
        </select>
	</td>
	
	<td width="117">Location:<br>
        <label for="DInsite"></label>
        <select name="DInsite" id="DInsite">
		  <option value=""></option>
	  
        <%
While (NOT DInsite.EOF)
%>
        <option value="<%=(DInsite.Fields.Item("equipment").Value)%>"<%If (Not isNull((equip_edit.Fields.Item("DInSite").Value))) Then If (CStr(DInsite.Fields.Item("equipment").Value) = CStr((equip_edit.Fields.Item("DInSite").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(DInsite.Fields.Item("equipment").Value)%></option>
        <%
  DInsite.MoveNext()
Wend
If (DInsite.CursorType > 0) Then
  DInsite.MoveFirst
Else
  DInsite.Requery
End If
%> 
        </select>
	</td>
	<td>Computer Name: <br>
      <label for="DInComputerName"></label>
      <input name="DInComputerName" type="text" id="DInComputerName" value="<%=(equip_edit.Fields.Item("DInComputerName").Value)%>" size="20" maxlength="40"/>
	</td>
  </tr>
  <tr>
	<td height = "5">
	</td>
  </tr>
  <tr>
	<td>
	</td>
	<td>Serial: <br>
      <label for="DInserialNum"></label>
      <input name="DInserialNum" type="text" id="DInserialNum" value="<%=(equip_edit.Fields.Item("DInserialNum").Value)%>" size="20" maxlength="50"/> 
	</td>
	<td>Asset Tag: <br>
      <label for="DInAssetTagNum"></label>
      <input name="DInAssetTagNum" type="text" id="DInAssetTagNum" value="<%=(equip_edit.Fields.Item("DInAssetTagNum").Value)%>" size="10" maxlength="40"/>
	</td>
	<td>Cube #: <br>
      <label for="DInCubeLocation"></label>
      <input name="DInCubeLocation" type="text" id="DInCubeLocation" value="<%=(equip_edit.Fields.Item("DInCubeLocation").Value)%>" size="10" maxlength="40"/>
	</td>
  </tr>
  </table>
  
  <img src="graphics/bar_break.png" width="800" height="10" />
  
  <table>
  <tr>
    <td>Ticket #: <br>
      <label for="DInticketnum"></label>
      <input name="DInticketnum" type="text" id="DInticketnum" value="<%=(equip_edit.Fields.Item("DInTicketNum").Value)%>" size="20" maxlength="20"/>  </td>
	<td>Assigned: <br>
      <label for="DInassigneduser"></label>
      <input name="DInassigneduser" type="text" id="DInassigneduser" value="<%=(equip_edit.Fields.Item("DInAssignedUser").Value)%>" size="20" maxlength="40"/> 
	</td>
	<td>User ID (AD ID): <br>
      <label for="DInUserID"></label>
      <input name="DInUserID" type="text" id="DInUserID" value="<%=(equip_edit.Fields.Item("DInUserID").Value)%>" size="20" maxlength="40"/>
	</td>
	<td>Cost Center: <br>
      <label for="DInCostCenter"></label>
      <input name="DInCostCenter" type="text" id="DInCostCenter" value="<%=(equip_edit.Fields.Item("DInCostCenter").Value)%>" size="10" maxlength="40"/>
	</td>
	<td>Archer Number: <br>
      <label for="DInArcherNum"></label>
      <input name="DInArcherNum" type="text" id="DInArcherNum" value="<%=(equip_edit.Fields.Item("DInArcherNum").Value)%>" size="10" maxlength="40"/>
	</td>
	</tr>
  
	<tr>
    <td>Deployed: <br>
      <label for="DIndeployeddate"></label>
      <input name="DIndeployeddate" type="text" id="DIndeployeddate" value="<%=(equip_edit.Fields.Item("DInDeployedDate").Value)%>" size="20" maxlength="40"/> 
	</td>
    <td>Returned Date: <br>
      <label for="DInenddate"></label>
      <input name="DInenddate" type="text" id="DInenddate" value="<%=(equip_edit.Fields.Item("DInEndDate").Value)%>" size="20" maxlength="40"/> 
	</td>



	
	</tr>
  
  
  
  
  <tr>

    <td colspan="5">Notes: <br>
      <label for="DInnotes"></label>
      <textarea name="DInnotes" cols="60" rows="4" type="text" id="DInnotes" maxlength="1500"><%=(equip_edit.Fields.Item("DInNotes").Value)%> </textarea></td>
    </tr>
  <tr align="center">
  
  <td height = "30"><input name="submit" type="submit" value="Update" /></td>
  </tr>
</table>
  <input type="hidden" name="MM_update" value="update" />
  <input type="hidden" name="MM_recordId" value="<%= equip_edit.Fields.Item("DInSerialNum").Value %>" />
</form></td>
  </tr>
  <tr>
  <td>
  </td>
  <td>
	<img src="graphics/bar_break.png" width="800" height="10" />
  </td>
  </tr>
</table>



<div id="mainbody">

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
