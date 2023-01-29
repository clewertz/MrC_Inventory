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
    MM_editCmd.CommandText = "INSERT INTO dbo.BF_Desktop (BF_Type,BF_Model,BF_Serial,BF_Status,BF_Notes,BF_Site,Image_Ver, Ticket_Num, Assigned, Deployed_Date, End_Date, BFComputerName) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("bftype")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 50, Request.Form("bfmodel")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("bfserial")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("bfstatus")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 1000, Request.Form("bfnotes")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, 50, Request.Form("bfsite")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, 50, Request.Form("imagevers")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 201, 1, 50, Request.Form("ticket_num")) ' adLongVarChar
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 201, 1, 50, Request.Form("assigned")) ' adLongVarChar	
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 135, 1, 10, MM_IIF(Request.Form("deployed"), Request.Form("deployed"), null)) ' adDBTimeStamp
	MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 135, 1, 10, MM_IIF(Request.Form("end_date"), Request.Form("end_date"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 201, 1, 50, Request.Form("bfcomputername")) ' adLongVarChar	
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
Dim bftype
Dim bftype_cmd
Dim bftype_numRows

Set bftype_cmd = Server.CreateObject ("ADODB.Command")
bftype_cmd.ActiveConnection = MM_NationStar_STRING
bftype_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '1'  ORDER BY equipment ASC" 
bftype_cmd.Prepared = true

Set bftype = bftype_cmd.Execute
bftype_numRows = 0
%>

<%
Dim bfmodel
Dim bfmodel_cmd
Dim bfmodel_numRows

Set bfmodel_cmd = Server.CreateObject ("ADODB.Command")
bfmodel_cmd.ActiveConnection = MM_NationStar_STRING
bfmodel_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '2'  ORDER BY equipment ASC" 
bfmodel_cmd.Prepared = true

Set bfmodel = bfmodel_cmd.Execute
bfmodel_numRows = 0
%>

<%
Dim imagevers
Dim imagevers_cmd
Dim imagevers_numRows

Set imagevers_cmd = Server.CreateObject ("ADODB.Command")
imagevers_cmd.ActiveConnection = MM_NationStar_STRING
imagevers_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '3'  ORDER BY equipment ASC" 
imagevers_cmd.Prepared = true

Set imagevers = imagevers_cmd.Execute
imagevers_numRows = 0
%>

<%
Dim bfstatus
Dim bfstatus_cmd
Dim bfstatus_numRows

Set bfstatus_cmd = Server.CreateObject ("ADODB.Command")
bfstatus_cmd.ActiveConnection = MM_NationStar_STRING
bfstatus_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '5'  ORDER BY equipment ASC" 
bfstatus_cmd.Prepared = true

Set bfstatus = bfstatus_cmd.Execute
bfstatus_numRows = 0
%>

<%
Dim bfsite
Dim bfsite_cmd
Dim bfsite_numRows

Set bfsite_cmd = Server.CreateObject ("ADODB.Command")
bfsite_cmd.ActiveConnection = MM_NationStar_STRING
bfsite_cmd.CommandText = "SELECT equipment FROM dbo.BF_Resource WHERE equipment_num = '4'  ORDER BY equipment ASC" 
bfsite_cmd.Prepared = true

Set bfsite = bfsite_cmd.Execute
bfsite_numRows = 0
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="citrix_migration_prep.css">
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
<title>Equip Add</title>
</head>

<body>

<div id = "header">Equipment Add</div>
<div id="menu" align="center"></div>
<br>

<table width="800" border="0">
  <tr>
    <td width="244">  
<!-- #include virtual= "nationstar/desktop/Library/sidebar.lbi" --></td>
    <td width="546"><form ACTION="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="300" border="0">
  <tr align="center">
    <td width="53">Type: <br>
            <select name="bftype" id="bftype">
        <option value=""></option>
        <%
While (NOT bftype.EOF)
%>
        <option value="<%=(bftype.Fields.Item("equipment").Value)%>"><%=(bftype.Fields.Item("equipment").Value)%></option>
        <%
  bftype.MoveNext()
Wend
If (bftype.CursorType > 0) Then
  bftype.MoveFirst
Else
  bftype.Requery
End If
%>

       
      </select>      <br></td>
    
	<td width="122">Model: <br>
      <label for="bfmodel"></label>
<select name="bfmodel" id="bfmodel">
	  <option value=""></option>
        <%
While (NOT bfmodel.EOF)
%>
        <option value="<%=(bfmodel.Fields.Item("equipment").Value)%>"><%=(bfmodel.Fields.Item("equipment").Value)%></option>
        <%
  bfmodel.MoveNext()
Wend
If (bfmodel.CursorType > 0) Then
  bfmodel.MoveFirst
Else
  bfmodel.Requery
End If
%>

       
      </select></td>
	  
	  <td width="117">Location:<br>
        <label for="bfsite"></label>
        <select name="bfsite" id="bfsite">
		  <option value=""></option>
	  
        <%
While (NOT bfsite.EOF)
%>
        <option value="<%=(bfsite.Fields.Item("equipment").Value)%>"><%=(bfsite.Fields.Item("equipment").Value)%></option>
        <%
  bfsite.MoveNext()
Wend
If (bfsite.CursorType > 0) Then
  bfsite.MoveFirst
Else
  bfsite.Requery
End If
%> 
        </select></td>
        
  <td width="105">Image Vers: <br>
       <label for="imagevers"></label>
      <select name="imagevers" id="imagevers">
		  <option value=""></option>
	  
        <%
While (NOT imagevers.EOF)
%>
        <option value="<%=(imagevers.Fields.Item("equipment").Value)%>"><%=(imagevers.Fields.Item("equipment").Value)%></option>
        <%
  imagevers.MoveNext()
Wend
If (imagevers.CursorType > 0) Then
  imagevers.MoveFirst
Else
  imagevers.Requery
End If
%> 
      </select> </td>      


		  <td width="95">Status:<br>
        <label for="bfstatus"></label>
        <select name="bfstatus" id="bfstatus">
		  <option value=""></option>
	  
        <%
While (NOT bfstatus.EOF)
%>
        <option value="<%=(bfstatus.Fields.Item("equipment").Value)%>"><%=(bfstatus.Fields.Item("equipment").Value)%></option>
        <%
  bfstatus.MoveNext()
Wend
If (bfstatus.CursorType > 0) Then
  bfstatus.MoveFirst
Else
  bfstatus.Requery
End If
%> 
        </select></td> 
        

	  
  </tr>
  <tr align="center">
    

	<td>Ticket #: <br>
      <label for="ticket_num"></label>
      <input name="ticket_num" type="text" id="ticket_num" size="20" maxlength="20"/>  </td>


    <td>Serial: <br>
      <label for="bfserial"></label>
      <input name="bfserial" type="text" id="bfserial" size="20" maxlength="20"/> 
	</td>
	  
	  
    <td>Assigned: <br>
      <label for="assigned"></label>
      <input name="assigned" type="text" id="assigned" size="20" maxlength="40"/> 
	</td>
    
    <td>Deployed: <br>
      <label for="deployed"></label>
      <input name="deployed" type="text" id="deployed" size="20" maxlength="40"/> 
	</td>
    
    <td>End Date: <br>
      <label for="end_date"></label>
      <input name="end_date" type="text" id="end_date" size="20" maxlength="40"/> 
	</td>
    
	
  </tr>
  <tr align="center">
    <td>Computer Name: <br>
      <label for="bfcomputername"></label>
      <input name="bfcomputername" type="text" id="bfcomputername" size="20" maxlength="40"/> 
	</td>
	
    <td colspan="4">Notes: <br>
      <label for="bfnotes"></label>
      <textarea name="bfnotes" cols="60" rows="4" type="text" id="bfnotes" maxlength="1500"> </textarea></td>
    </tr>
  <tr align="center">
  <td></td>
  <td><input name="submit" type="submit" value="Submit" /></td>
  <td><input name="Reset" type="reset" value="Reset" /></td>
  </tr>
</table>
  <input type="hidden" name="MM_insert" value="form1" />
</form></td>
  </tr>
</table>



<div id="mainbody">

</body>
</html>

<%
bfmodel.Close()
Set bfmodel = Nothing
%>

<%
imagevers.Close()
Set imagevers = Nothing
%>

<%
bfstatus.Close()
Set bfstatus = Nothing
%>

<%
bfsite.Close()
Set bfsite = Nothing
%>
